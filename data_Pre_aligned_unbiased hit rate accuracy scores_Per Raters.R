# 加载必要的R包
library(dplyr)
library(tidyr)
library(readxl)
library(openxlsx)

# Step 1: 读取数据
file_path <- "aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型
data <- data %>%
  mutate(Expression_Type = case_when(
    grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
    grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
    grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
    grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
    grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
    TRUE ~ "Other"
  ))

# Step 3: 将 Categorizing_Expressions_Score 映射为 Chosen_Expression
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

# 初始化存储结果的数据框
results <- data.frame(
  CASE = integer(),
  Expression_Type = character(),
  UHR = numeric(),
  Chance_UHR = numeric(),
  Performance_Above_Chance = numeric(),
  stringsAsFactors = FALSE
)

# 初始化存储不完整数据的CASE列表
incomplete_cases <- list()

# 初始化Excel工作簿
wb <- createWorkbook()

# 计算 UHR 和 Chance UHR
calculate_uhr_chance_uhr <- function(conf_matrix) {
  uhr_results <- numeric(length(rownames(conf_matrix)))
  chance_uhr_results <- numeric(length(rownames(conf_matrix)))
  N <- sum(conf_matrix)  # 总和 N
  
  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]  # 对角线元素
    b <- sum(conf_matrix[emotion, ]) - a  # 行中的非对角线元素之和
    d <- sum(conf_matrix[, emotion]) - a  # 列中的非对角线元素之和
    
    # 检查 b 和 d 是否为0，避免除以零
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

# Step 4: 按CASE计算 Unbiased_Hit_Rate 和 Chance_Unbiased_Hit_Rate
cases <- unique(data$CASE)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

for (case in cases) {
  # 筛选出当前CASE的数据
  case_data <- data %>% filter(CASE == case)
  
  # 创建选择矩阵，排除 "Other" 类别
  choice_matrix <- case_data %>%
    filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
    count(Expression_Type, Chosen_Expression) %>%
    spread(key = Chosen_Expression, value = n, fill = 0) %>%
    complete(Expression_Type = expected_emotions, fill = list(n = 0))
  
  # 转换为普通数据框以支持行名
  choice_matrix <- as.data.frame(choice_matrix)
  
  # 检查是否包含所有预期的情绪类别
  missing_emotions <- setdiff(expected_emotions, colnames(choice_matrix))
  
  if (length(missing_emotions) > 0) {
    # 如果有缺失情绪类别，记录CASE并跳过此CASE的计算
    incomplete_cases[[length(incomplete_cases) + 1]] <- list(CASE = case, Missing_Emotions = missing_emotions)
    next
  }
  
  rownames(choice_matrix) <- choice_matrix$Expression_Type
  choice_matrix <- choice_matrix[, -1]
  
  # 将数据框转换为矩阵
  choice_matrix <- as.matrix(choice_matrix)
  
  # 将混淆矩阵保存到Excel工作簿
  addWorksheet(wb, paste("CASE", case))
  writeData(wb, paste("CASE", case), choice_matrix, rowNames = TRUE)
  
  # 计算 UHR 和 Chance UHR
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

# 保存混淆矩阵到文件
saveWorkbook(wb, "confusion_matrices_all_cases.xlsx", overwrite = TRUE)

# 打印缺失情绪类别的CASE
if (length(incomplete_cases) > 0) {
  cat("Incomplete CASEs:\n")
  for (case_info in incomplete_cases) {
    cat("CASE:", case_info$CASE, "- Missing Emotions:", paste(case_info$Missing_Emotions, collapse = ", "), "\n")
  }
}

# 保存 UHR 和 Chance UHR 结果到 Excel 文件
write.xlsx(results, "uhr_results_all_cases.xlsx", rowNames = FALSE)

# 打印最终结果
print(results)

# 安装并加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")
install.packages("writexl")

library(readxl)
library(dplyr)
library(tidyr)
library(writexl)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型
data <- data %>%
  mutate(Expression_Type = case_when(
    grepl("dis", Material) ~ "Disgust",
    grepl("enj", Material) ~ "Enjoyment",
    grepl("aff", Material) ~ "Affiliation",
    grepl("dom", Material) ~ "Dominance",
    grepl("neu", Material) ~ "Neutral",
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
    TRUE ~ NA_character_  # 如果遇到未匹配的值，将其设为 NA
  ))

# 初始化存储结果的数据框
results <- data.frame(
  CASE = integer(),
  Expression_Type = character(),
  Unbiased_Hit_Rate = numeric(),
  Chance_Unbiased_Hit_Rate = numeric(),
  Performance_Above_Chance = numeric(),
  stringsAsFactors = FALSE
)

# Step 4: 按CASE计算Unbiased_Hit_Rate和Chance_Unbiased_Hit_Rate
cases <- unique(data$CASE)

for (case in cases) {
  # 筛选出当前CASE的数据
  case_data <- data %>% filter(CASE == case)
  
  # 创建选择矩阵
  choice_matrix <- case_data %>%
    count(Expression_Type, Chosen_Expression) %>%
    spread(Chosen_Expression, n, fill = 0)
  
  # 计算行边际值和列边际值
  row_margins <- rowSums(choice_matrix[,-1])
  col_margins <- colSums(choice_matrix[,-1])
  
  # 计算 UHR，包括 "Other" 对其他表情的影响
  uhr_values <- mapply(function(row, col, value, row_type, col_type) {
    if (row_type == "Other" || col_type == "Other") {
      return(0)  # 假设 "Other" 的行和列对最终 UHR 不贡献直接值
    } else if (row == 0 || col == 0) {
      return(0)
    } else {
      return((value^2) / (row * col))
    }
  }, 
  row = rep(row_margins, times = length(col_margins)), 
  col = rep(col_margins, each = length(row_margins)), 
  value = as.vector(as.matrix(choice_matrix[,-1])),
  row_type = rep(choice_matrix$Expression_Type, times = length(col_margins)),
  col_type = rep(names(choice_matrix)[-1], each = length(row_margins)))
  
  # 将 UHR 值转为矩阵形式并计算最终 UHR
  uhr_matrix <- matrix(uhr_values, nrow = nrow(choice_matrix), ncol = ncol(choice_matrix) - 1, byrow = TRUE)
  uhr_final <- rowSums(uhr_matrix)
  
  # 过滤掉 "Other" 作为实际表情类型，并显示结果
  uhr_summary <- data.frame(Expression_Type = choice_matrix$Expression_Type, Unbiased_Hit_Rate = uhr_final)
  uhr_summary <- uhr_summary %>% filter(Expression_Type != "Other")
  
  # 计算每个CASE的期望的机会水平 UHR
  expression_probs <- case_data %>%
    count(Expression_Type) %>%
    mutate(Prob_Expression_Type = n / sum(n))
  
  chosen_expression_probs <- case_data %>%
    count(Chosen_Expression) %>%
    mutate(Prob_Chosen_Expression = n / sum(n))
  
  expected_probs <- expression_probs %>%
    full_join(chosen_expression_probs, by = c("Expression_Type" = "Chosen_Expression")) %>%
    mutate(Chance_Unbiased_Hit_Rate = Prob_Expression_Type * Prob_Chosen_Expression) %>%
    select(Expression_Type, Chance_Unbiased_Hit_Rate)
  
  # 计算每个CASE的Performance_Above_Chance
  uhr_comparison <- uhr_summary %>%
    left_join(expected_probs, by = "Expression_Type") %>%
    mutate(
      CASE = case,
      Performance_Above_Chance = Unbiased_Hit_Rate - Chance_Unbiased_Hit_Rate
    )
  
  # 存储结果
  results <- bind_rows(results, uhr_comparison)
}

# Step 5: 将结果按CASE分开显示，并保存为CSV文件
results_by_case <- results %>%
  pivot_wider(names_from = Expression_Type, values_from = c(Unbiased_Hit_Rate, Chance_Unbiased_Hit_Rate, Performance_Above_Chance))

# 保存按CASE分开的结果为CSV文件
write.csv(results_by_case, "UHR_and_Chance_UHR_by_CASE.csv", row.names = FALSE)

# 输出结果
print("Summary of UHR and Chance-Level UHR by CASE:")
print(results_by_case)

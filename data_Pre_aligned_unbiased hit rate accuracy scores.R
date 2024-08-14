# 加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")
install.packages("car")  # 用于 ANOVA
install.packages("emmeans")  # 用于事后检验
install.packages("writexl")  # 用于保存结果

library(readxl)
library(dplyr)
library(tidyr)
library(car)
library(emmeans)
library(writexl)
library(ggplot2)  # 用于可视化

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

# Step 4: 创建选择矩阵
choice_matrix <- data %>%
  count(Expression_Type, Chosen_Expression) %>%
  spread(Chosen_Expression, n, fill = 0)

# Step 5: 计算行边际值和列边际值
row_margins <- rowSums(choice_matrix[,-1])
col_margins <- colSums(choice_matrix[,-1])

# Step 6: 计算 UHR，包括 "Other" 对其他表情的影响
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

# Step 7: 过滤掉 "Other" 作为实际表情类型，并显示结果
uhr_summary <- data.frame(Expression_Type = choice_matrix$Expression_Type, Unbiased_Hit_Rate = uhr_final)
uhr_summary <- uhr_summary %>% filter(Expression_Type != "Other")

# 打印 UHR 结果
print(uhr_summary)

# Step 8: Arcsine transformation of UHR
uhr_summary <- uhr_summary %>%
  mutate(Arcsine_UHR = asin(sqrt(Unbiased_Hit_Rate)))

# Step 9: 为重复测量 ANOVA 准备数据
# 使用 CASE 列作为参与者ID
participant_data <- data %>%
  left_join(uhr_summary, by = "Expression_Type") %>%
  group_by(CASE, Expression_Type) %>%
  summarise(Arcsine_UHR = mean(Arcsine_UHR, na.rm = TRUE), .groups = 'drop')

# Step 10: 进行重复测量 ANOVA
aov_results <- aov(Arcsine_UHR ~ Expression_Type + Error(CASE/Expression_Type), data = participant_data)

# 查看ANOVA结果
summary(aov_results)

# Step 10: 进行重复测量 ANOVA
aov_results <- aov(Arcsine_UHR ~ Expression_Type + Error(CASE/Expression_Type), data = participant_data)

# Step 11: 进行事后检验 (例如 emmeans 包)
emms <- emmeans(aov_results, ~ Expression_Type)
pairwise_comparison <- pairs(emms)

# 保存 ANOVA 结果和事后检验结果为同一个文本文件
capture.output({
  cat("ANOVA Results:\n")
  print(summary(aov_results))
  cat("\n\nPairwise Comparisons (Post Hoc Test):\n")
  print(pairwise_comparison)
}, file = "analysis_ANOVA_UHR.txt")

# Step 12: 可视化 UHR 方差分析的结果
mean_uhr <- participant_data %>%
  group_by(Expression_Type) %>%
  summarise(
    Mean_UHR = mean(Arcsine_UHR, na.rm = TRUE),
    SE = sd(Arcsine_UHR, na.rm = TRUE) / sqrt(n())
  )

ggplot(mean_uhr, aes(x = Expression_Type, y = Mean_UHR)) +
  geom_line(group = 1, color = "blue") +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = Mean_UHR - SE, ymax = Mean_UHR + SE), width = 0.2) +
  labs(title = "Mean Arcsine UHR by Expression Type",
       x = "Expression Type",
       y = "Mean Arcsine UHR") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 10))

# Step 13: 手动构建混淆矩阵以检验
confusion_matrix <- data %>%
  count(Expression_Type, Chosen_Expression) %>%
  spread(Chosen_Expression, n, fill = 0)

# Add row and column totals
confusion_matrix <- confusion_matrix %>%
  mutate(Total = rowSums(.[-1])) %>%
  add_row(Expression_Type = "Total", !!!colSums(confusion_matrix[-1]))

# Print the confusion matrix
print(confusion_matrix)

# Step 14: 比较实际 UHR 与期望 UHR：计算期望的机会水平 UHR，并与观察到的 UHR 进行比较
expression_probs <- data %>%
  count(Expression_Type) %>%
  mutate(Prob_Expression_Type = n / sum(n))

chosen_expression_probs <- data %>%
  count(Chosen_Expression) %>%
  mutate(Prob_Chosen_Expression = n / sum(n))

expected_probs <- expression_probs %>%
  full_join(chosen_expression_probs, by = c("Expression_Type" = "Chosen_Expression")) %>%
  mutate(Chance_Unbiased_Hit_Rate = Prob_Expression_Type * Prob_Chosen_Expression) %>%
  select(Expression_Type, Chance_Unbiased_Hit_Rate)

uhr_comparison <- uhr_summary %>%
  left_join(expected_probs, by = "Expression_Type") %>%
  mutate(
    Performance_Above_Chance = Unbiased_Hit_Rate - Chance_Unbiased_Hit_Rate
  )

# Print comparison
print("Comparison of observed UHR with expected chance-level UHR:")
print(uhr_comparison)

# 保存比较结果为CSV文件
write.csv(uhr_comparison, "UHR_Comparison_with_Chance.csv", row.names = FALSE)

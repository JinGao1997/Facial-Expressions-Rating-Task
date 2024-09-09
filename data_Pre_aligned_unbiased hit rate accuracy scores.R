# 加载必要的R包
library(dplyr)
library(tidyr)
library(readxl)
library(openxlsx)
library(car)  # 用于ANOVA
library(emmeans)  # 用于事后检验
library(ggplot2)  # 可视化
library(ggpubr)   # Q-Q 图

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

# Step 4: 定义 UHR 和 Chance UHR 计算函数
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

# 初始化存储不完整数据的CASE列表
incomplete_cases <- list()

# Step 5: 按 CASE 计算 UHR 和 Chance UHR
cases <- unique(data$CASE)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")
# 检查每个参与者的选择矩阵是否包含所有预期的情绪类别
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
  
  # 检查是否包含所有预期的表情类别
  missing_emotions <- setdiff(expected_emotions, colnames(choice_matrix))
  
  # 如果有缺失情绪类别，记录CASE并跳过此CASE的计算
  if (length(missing_emotions) > 0) {
    incomplete_cases[[length(incomplete_cases) + 1]] <- list(CASE = case, Missing_Emotions = missing_emotions)
    cat("Warning: Missing emotions in CASE", case, "- Missing:", missing_emotions, "\n")
    next
  }
  
  # 正常进行 UHR 和 Chance UHR 计算
  rownames(choice_matrix) <- choice_matrix$Expression_Type
  choice_matrix <- choice_matrix[, -1]
  
  # 将数据框转换为矩阵
  choice_matrix <- as.matrix(choice_matrix)
  
  # 打印选择矩阵以确认其内容
  print(choice_matrix)

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


# 打印结果
print(results)

# Step 6: Arcsine transformation of UHR
results <- results %>%
  mutate(Arcsine_UHR = asin(sqrt(UHR)))

# Step 7: 计算不同 Expression_Type 的 UHR 和 Chance UHR 平均数和标准差
uhr_stats <- results %>%
  group_by(Expression_Type) %>%
  summarise(
    Mean_UHR = mean(UHR, na.rm = TRUE),
    SD_UHR = sd(UHR, na.rm = TRUE),
    Mean_Chance_UHR = mean(Chance_UHR, na.rm = TRUE),
    SD_Chance_UHR = sd(Chance_UHR, na.rm = TRUE)
  )

# 保存 UHR 和 Chance UHR 的平均数和标准差到 CSV 文件
write.csv(uhr_stats, "UHR_ChanceUHR_Avg_SD.csv", row.names = FALSE)
# Step 7: 进行重复测量 ANOVA
aov_results <- aov(Arcsine_UHR ~ Expression_Type + Error(CASE/Expression_Type), data = results)

# 查看 ANOVA 结果
summary(aov_results)

# Step 8: 进行事后检验
emms <- emmeans(aov_results, ~ Expression_Type)
pairwise_comparison <- pairs(emms)

# 保存 ANOVA 结果和事后检验结果
capture.output({
  cat("ANOVA Results:\n")
  print(summary(aov_results))
  cat("\n\nPairwise Comparisons (Post Hoc Test):\n")
  print(pairwise_comparison)
}, file = "analysis_ANOVA_UHR.txt")

# Step 9: 可视化结果
mean_uhr <- results %>%
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

# 保存结果
write.xlsx(results, "uhr_results_all_cases.xlsx", rowNames = FALSE)

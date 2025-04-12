# ---------------------------------------------------------------------------
# 加载必要的 R 包
library(dplyr)
library(tidyr)
library(readxl)
library(openxlsx)
library(car)         # 可选，用于其他分析
library(emmeans)     # 用于事后检验
library(ggplot2)     # 用于可视化
library(ggpubr)      # 用于 Q-Q 图等
library(lmerTest)    # 用于混合效应模型分析
library(stringr)     # 用于字符串处理

# ---------------------------------------------------------------------------
# Step 1: 读取数据 / Read data
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# ---------------------------------------------------------------------------
# Step 1.1: 转换评分者性别 / Convert Rater_Gender from raw Gender column
# 注意：原始数据中 Gender 列的 1 表示女性，2 表示男性
data <- data %>%
  mutate(Rater_Gender = case_when(
    Gender == 1 ~ "Female",
    Gender == 2 ~ "Male",
    TRUE ~ NA_character_
  ))

# ---------------------------------------------------------------------------
# 【调试提示】（可选）检查 Material 字段中是否包含 "male" 和 "fema"
# debug_materials <- data %>%
#   mutate(has_male = grepl("male", Material, ignore.case = TRUE),
#          has_fema = grepl("fema", Material, ignore.case = TRUE)) %>%
#   select(Material, has_male, has_fema)
# print(head(debug_materials, 20))

# ---------------------------------------------------------------------------
# Step 2: 从 Material 提取表情类别 / Extract Expression_Type from Material
data <- data %>%
  mutate(Expression_Type = case_when(
    grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
    grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
    grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
    grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
    grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
    TRUE ~ "Other"
  ))

# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Step 3.1: 从 Material 提取其他变量
# Face_Gender：先匹配 "fema"（例如 R_enjFema18），再匹配 "male"（例如 L_domMale89 或 L_affMale41）
# Version：根据 Material 中是否包含 "L" 或 "R"
# 注意：Group 直接使用原始数据中的 Group 变量（表示被试被分配到的实验设计组）
data <- data %>%
  mutate(
    Face_Gender = case_when(
      grepl("fema", Material, ignore.case = TRUE) ~ "Female",  
      grepl("male", Material, ignore.case = TRUE) ~ "Male",
      TRUE ~ NA_character_
    ),
    Version = case_when(
      grepl("L", Material, ignore.case = TRUE) ~ "L",
      grepl("R", Material, ignore.case = TRUE) ~ "R",
      TRUE ~ NA_character_
    )
    # Group 直接沿用原始数据中的 Group 变量
  )

# ---------------------------------------------------------------------------
# Step 4: 初始化存储结果的数据框 / Initialize a dataframe to store results
# 这里我们按 CASE、Expression_Type 以及 Face_Gender 分组保存计算结果
results <- data.frame(
  CASE = integer(),
  Expression_Type = character(),
  Face_Gender = character(),
  UHR = numeric(),
  Chance_UHR = numeric(),
  Performance_Above_Chance = numeric(),
  Rater_Gender = character(),
  Version = character(),
  Group = character(),  # 使用原始数据中的 Group 信息
  stringsAsFactors = FALSE
)

# ---------------------------------------------------------------------------
# Step 5: 定义 UHR 和 Chance UHR 的计算函数
calculate_uhr_chance_uhr <- function(conf_matrix) {
  uhr_results <- numeric(nrow(conf_matrix))
  chance_uhr_results <- numeric(nrow(conf_matrix))
  # 总计数 N 忽略 NA
  N <- sum(conf_matrix, na.rm = TRUE)
  
  for (i in seq_len(nrow(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]
    b <- sum(conf_matrix[emotion, ], na.rm = TRUE) - a
    d <- sum(conf_matrix[, emotion], na.rm = TRUE) - a
    
    # 如果有 NA，则直接设为 0；否则，如果 (a+b) 和 (a+d) > 0 且 N > 0，则计算 UHR
    if (is.na(a) || is.na(b) || is.na(d) || is.na(N)) {
      UHR <- 0
      Pc <- 0
    } else if ((a + b) > 0 && (a + d) > 0 && N > 0) {
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

# ---------------------------------------------------------------------------
# 初始化存储不完整数据的 CASE 列表（调试用）
incomplete_cases <- list()

# ---------------------------------------------------------------------------
# Step 6: 按 CASE 与 Face_Gender 构建完整混淆矩阵，然后计算 UHR 和 Chance UHR

# 定义预期的表情水平（确保所有期望的表达都在此列表中）
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

cases <- unique(data$CASE)

for (case in cases) {
  # 筛选当前 CASE 的数据
  case_data <- data %>% filter(CASE == case)
  
  # 获取当前 CASE 中存在的所有 Face_Gender
  face_genders <- unique(case_data$Face_Gender)
  
  for (fg in face_genders) {
    # 筛选该 CASE 中当前 Face_Gender 的所有记录
    fg_data <- case_data %>% filter(Face_Gender == fg)
    
    # 仅考虑预期表达和 Chosen_Expression 不为 "Other" 的记录
    subset_data <- fg_data %>% filter(Expression_Type %in% expected_emotions,
                                      Chosen_Expression != "Other")
    if(nrow(subset_data) == 0) next
    
    # 构造完整的混淆矩阵：统计 Expression_Type 与 Chosen_Expression 的频数，
    # 并确保所有预期表达都作为行和列显示（缺失值填 0）
    choice_matrix <- subset_data %>%
      count(Expression_Type, Chosen_Expression) %>%
      complete(
        Expression_Type = factor(expected_emotions, levels = expected_emotions),
        Chosen_Expression = factor(expected_emotions, levels = expected_emotions),
        fill = list(n = 0)
      ) %>%
      spread(key = Chosen_Expression, value = n, fill = 0)
    
    # 将 NA 替换为 0（保险起见）
    choice_matrix[is.na(choice_matrix)] <- 0
    
    # 转换为矩阵
    choice_matrix <- as.data.frame(choice_matrix)
    rownames(choice_matrix) <- choice_matrix$Expression_Type
    choice_matrix <- choice_matrix[, -1]
    choice_matrix <- as.matrix(choice_matrix)
    
    # 计算 UHR 和 Chance UHR（调用预先定义的函数）
    uhr_chance_results <- calculate_uhr_chance_uhr(conf_matrix = choice_matrix)
    
    # 记录当前 CASE + Face_Gender 下，每个预期表情的计算结果
    for (em in expected_emotions) {
      # 确保该表情存在于混淆矩阵中（因为 complete() 应该保证这一点）
      if (!(em %in% rownames(choice_matrix))) next
      
      # 取该子集中的第一条记录作为其他变量的代表
      Rater_Gender_val <- subset_data$Rater_Gender[1]
      Version_val <- subset_data$Version[1]
      Group_val <- as.character(subset_data$Group[1])  # 使用原始数据中的 Group
      
      # 将结果添加到 results 数据框中
      results <- rbind(results, data.frame(
        CASE = case,
        Expression_Type = em,
        Face_Gender = fg,
        UHR = uhr_chance_results$UHR[which(rownames(choice_matrix) == em)],
        Chance_UHR = uhr_chance_results$Chance_UHR[which(rownames(choice_matrix) == em)],
        Performance_Above_Chance = uhr_chance_results$UHR[which(rownames(choice_matrix) == em)] - 
          uhr_chance_results$Chance_UHR[which(rownames(choice_matrix) == em)],
        Rater_Gender = Rater_Gender_val,
        Version = Version_val,
        Group = Group_val,
        stringsAsFactors = FALSE
      ))
    }
  }
}

# 调试用：打印部分结果，并查看 Face_Gender 水平
print(results)
cat("Face_Gender levels in results:", levels(factor(results$Face_Gender)), "\n")


# ---------------------------------------------------------------------------
# Step 7: 对 UHR 进行 Arcsine 转换
results <- results %>%
  mutate(Arcsine_UHR = asin(sqrt(UHR)))

# 将相关变量转换为因子，便于后续建模
results <- results %>%
  mutate(
    Face_Gender = factor(Face_Gender),
    Rater_Gender = factor(Rater_Gender),
    # 设置 Expression_Type 的因子水平，并将参考水平设为 "Neutral"
    Expression_Type = factor(Expression_Type, levels = c("Neutral", "Affiliation", "Disgust", "Dominance", "Enjoyment")),
    Version = factor(Version),
    Group = factor(Group)
  )

# 可选：统计各 Expression_Type 的均值和标准差，并保存到 CSV 文件
uhr_stats <- results %>%
  group_by(Expression_Type) %>%
  summarise(
    Mean_UHR = mean(UHR, na.rm = TRUE),
    SD_UHR = sd(UHR, na.rm = TRUE),
    Mean_Chance_UHR = mean(Chance_UHR, na.rm = TRUE),
    SD_Chance_UHR = sd(Chance_UHR, na.rm = TRUE)
  )
write.csv(uhr_stats, "UHR_ChanceUHR_Avg_SD.csv", row.names = FALSE)
cat("UHR 及 Chance UHR 的统计结果已保存为 UHR_ChanceUHR_Avg_SD.csv。\n")

# ---------------------------------------------------------------------------
# Step 8: 构建混合效应模型（LMM）
lmer_model <- lmer(Arcsine_UHR ~ Face_Gender + Rater_Gender + Expression_Type + Version + Group + (1|CASE), 
                   data = results)
summary(lmer_model)

# 使用 emmeans 对 Expression_Type 进行两两事后检验
emms <- emmeans(lmer_model, ~ Expression_Type)
pairwise_comparison <- pairs(emms)

# 将模型及事后检验结果保存到文本文件中
capture.output({
  cat("Mixed Effects Model Results:\n")
  print(summary(lmer_model))
  cat("\n\nPairwise Comparisons (Post Hoc Test):\n")
  print(pairwise_comparison)
}, file = "analysis_LMM_UHR.txt")
cat("混合效应模型及事后检验结果已保存为 analysis_LMM_UHR.txt。\n")

# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Step 10: 保存最终结果到 Excel 文件
write.xlsx(results, "uhr_results_all_cases.xlsx", rowNames = FALSE)
cat("最终结果已保存为 uhr_results_all_cases.xlsx。\n")

# ---------------------------------------------------------------------------
# Step 11: 检查问题案例（可选）
valid_choices <- data %>%
  filter(Expression_Type %in% expected_emotions, Chosen_Expression != "Other")
problematic_cases <- valid_choices %>%
  group_by(CASE, Expression_Type) %>%
  summarise(
    total_choices = n(),
    distinct_genders = n_distinct(Face_Gender),
    genders = paste(unique(Face_Gender), collapse = ", "),
    .groups = "drop"
  ) %>%
  filter(total_choices == 1 | distinct_genders == 1)
print(problematic_cases)

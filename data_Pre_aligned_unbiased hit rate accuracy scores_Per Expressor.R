# 加载必要的R包 / Load necessary packages
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)

# Step 1: 读取数据 / Read data
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型、Expressor 和性别（Gender）/ Extract Expression_Type, Expressor, and Gender from Material column
data <- data %>%
  mutate(
    Expression_Type = case_when(
      grepl("dis", Material) ~ "Disgust",
      grepl("enj", Material) ~ "Enjoyment",
      grepl("aff", Material) ~ "Affiliation",
      grepl("dom", Material) ~ "Dominance",
      grepl("neu", Material) ~ "Neutral",
      TRUE ~ "Other"
    ),
    Expressor = str_extract(Material, "Fema\\d+|Male\\d+"),
    Gender = case_when(
      str_detect(Material, "Fema") ~ "Female",
      str_detect(Material, "Male") ~ "Male"
    )
  )

# Step 3: 将 Categorizing_Expressions_Score 映射为 Chosen_Expression / Map Categorizing_Expressions_Score to Chosen_Expression
data <- data %>%
  mutate(
    Chosen_Expression = case_when(
      Categorizing_Expressions_Score == 1 ~ "Enjoyment",
      Categorizing_Expressions_Score == 2 ~ "Affiliation",
      Categorizing_Expressions_Score == 3 ~ "Dominance",
      Categorizing_Expressions_Score == 4 ~ "Disgust",
      Categorizing_Expressions_Score == 5 ~ "Neutral",
      Categorizing_Expressions_Score == 6 ~ "Other",
      TRUE ~ NA_character_
    )
  )

# 初始化存储结果的数据框 / Initialize a dataframe to store results
uhr_results_all <- data.frame(
  Expressor = character(),
  Expression_Type = character(),
  UHR = numeric(),
  Chance_UHR = numeric(),
  stringsAsFactors = FALSE
)

# Step 4: 定义 UHR 和 Chance UHR 计算函数 / Define functions to calculate UHR and Chance UHR
calculate_uhr_chance_uhr <- function(conf_matrix) {
  uhr_results <- numeric(length(rownames(conf_matrix)))
  chance_uhr_results <- numeric(length(rownames(conf_matrix)))
  N <- sum(conf_matrix)  # 总和 N / Total sum N

  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]  # 对角线元素 / Diagonal element
    b <- sum(conf_matrix[emotion, ]) - a  # 行中的非对角线元素之和 / Sum of non-diagonal elements in the row
    d <- sum(conf_matrix[, emotion]) - a  # 列中的非对角线元素之和 / Sum of non-diagonal elements in the column

    # 检查是否避免除以零 / Check to avoid division by zero
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

# Step 5: 按 Expressor 计算 UHR 和 Chance UHR / Calculate UHR and Chance UHR by Expressor
expressors <- unique(data$Expressor)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

for (expr in expressors) {
  # 筛选出当前 Expressor 的数据 / Filter data for the current Expressor
  expr_data <- data %>% filter(Expressor == expr)

  # 创建混淆矩阵，排除 "Other" 类别 / Create confusion matrix, excluding "Other" category
  conf_matrix <- expr_data %>%
    filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
    count(Expression_Type, Chosen_Expression) %>%
    tidyr::spread(key = Chosen_Expression, value = n, fill = 0) %>%
    tidyr::complete(Expression_Type = expected_emotions, fill = list(n = 0))

  # 转换为数据框并设置行名 / Convert to dataframe and set row names
  conf_matrix <- as.data.frame(conf_matrix)
  rownames(conf_matrix) <- conf_matrix$Expression_Type
  conf_matrix <- conf_matrix[, -1]

  # 确保所有预期的表情类别都存在于列中 / Ensure all expected emotions are in the columns
  missing_cols <- setdiff(expected_emotions, colnames(conf_matrix))
  for (col in missing_cols) {
    conf_matrix[[col]] <- 0
  }

  # 重新排序列以匹配预期的表情类别 / Reorder columns to match expected emotions
  conf_matrix <- conf_matrix[, expected_emotions]

  # 转换为矩阵 / Convert to matrix
  conf_matrix <- as.matrix(conf_matrix)

  # 计算 UHR 和 Chance UHR / Calculate UHR and Chance UHR
  uhr_chance_results <- calculate_uhr_chance_uhr(conf_matrix)

  # 收集结果 / Collect results
  for (i in seq_along(expected_emotions)) {
    uhr_results_all <- rbind(uhr_results_all, data.frame(
      Expressor = expr,
      Expression_Type = expected_emotions[i],
      UHR = uhr_chance_results$UHR[i],
      Chance_UHR = uhr_chance_results$Chance_UHR[i]
    ))
  }
}

# Step 6: 计算 Performance_Above_Chance / Calculate Performance Above Chance
uhr_results_all <- uhr_results_all %>%
  mutate(Performance_Above_Chance = UHR - Chance_UHR)

# Step 7: 添加 Gender 信息 / Add Gender information
expressor_gender <- data %>%
  dplyr::select(Expressor, Gender) %>%
  distinct()

uhr_results_all <- uhr_results_all %>%
  left_join(expressor_gender, by = "Expressor")

# Step 8: 保存每个 Expressor 的 UHR 信息 / Save UHR information for each Expressor
write.csv(uhr_results_all, "UHR_by_Expressor_and_Emotion.csv", row.names = FALSE)

# Step 9: 可视化 UHR 数据 / Visualize UHR data
uhr_by_expressor <- uhr_results_all %>%
  group_by(Expressor, Gender, Expression_Type) %>%
  summarise(Average_UHR = mean(UHR, na.rm = TRUE), .groups = 'drop')

# Step 10: 将表情类型缩写为前三个字母 / Abbreviate Expression_Type to first three letters
uhr_by_expressor <- uhr_by_expressor %>%
  mutate(Expression_Type = substr(Expression_Type, 1, 3))

# 绘制图形 / Plot the graph
ggplot(uhr_by_expressor, aes(x = Expression_Type, y = Average_UHR, fill = Gender)) +
  geom_bar(stat = "identity", position = "dodge") +
  facet_wrap(~ Expressor) +
  labs(title = "Average UHR by Expressor and Expression Type",
       x = "Expression Type",
       y = "Average UHR",
       fill = "Gender") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 60, hjust = 1, size = 5),
    axis.title.x = element_text(size = 10),
    axis.title.y = element_text(size = 8),
    plot.title = element_text(size = 12, face = "bold"),
    strip.text = element_text(size = 8)
  )

# Step 11: 按性别分组并排序 Expressors / Group and sort Expressors by Gender
uhr_sorted_female <- uhr_by_expressor %>%
  filter(Gender == "Female") %>%
  group_by(Expressor) %>%
  summarise(Average_UHR = mean(Average_UHR), .groups = 'drop') %>%
  arrange(desc(Average_UHR))

uhr_sorted_male <- uhr_by_expressor %>%
  filter(Gender == "Male") %>%
  group_by(Expressor) %>%
  summarise(Average_UHR = mean(Average_UHR), .groups = 'drop') %>%
  arrange(desc(Average_UHR))

# 检查最终排序结果 / Check final sorted results
print("按平均 UHR 排序的 Female Expressors：")
print(uhr_sorted_female)

print("按平均 UHR 排序的 Male Expressors：")
print(uhr_sorted_male)

# Step 12: 将排序后的结果保存为 CSV 文件 / Save sorted results to CSV files
write.csv(uhr_sorted_female, "Sorted_Female_Expressors_by_UHR.csv", row.names = FALSE)
write.csv(uhr_sorted_male, "Sorted_Male_Expressors_by_UHR.csv", row.names = FALSE)

# Step 13: 创建 Male29 和 Male75 的选择矩阵 / Create choice matrices for Male29 and Male75
Male29_data <- data %>%
  filter(Expressor == "Male29")

Male75_data <- data %>%
  filter(Expressor == "Male75")

# 检查数据是否存在 / Check if data exists
if (nrow(Male29_data) == 0) {
  stop("No data found for Male29. Please check the Expressor name or data.")
}

if (nrow(Male75_data) == 0) {
  stop("No data found for Male75. Please check the Expressor name or data.")
}

# 生成混淆矩阵 / Generate confusion matrices
Male29_matrix <- Male29_data %>%
  filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
  count(Expression_Type, Chosen_Expression) %>%
  tidyr::spread(Chosen_Expression, n, fill = 0)

Male75_matrix <- Male75_data %>%
  filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
  count(Expression_Type, Chosen_Expression) %>%
  tidyr::spread(Chosen_Expression, n, fill = 0)

# 将 Expression_Type 作为行名 / Set Expression_Type as row names
rownames(Male29_matrix) <- Male29_matrix$Expression_Type
Male29_matrix <- Male29_matrix[, -1]

rownames(Male75_matrix) <- Male75_matrix$Expression_Type
Male75_matrix <- Male75_matrix[, -1]

# Step 14: 可视化 / Visualization
Male29_long <- Male29_data %>%
  filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
  count(Expression_Type, Chosen_Expression) %>%
  ungroup()

Male75_long <- Male75_data %>%
  filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
  count(Expression_Type, Chosen_Expression) %>%
  ungroup()

# 添加 Expressor 信息 / Add Expressor information
Male29_long$Expressor <- "Male29"
Male75_long$Expressor <- "Male75"

# 合并数据 / Combine data
combined_data <- bind_rows(Male29_long, Male75_long)

# 绘制分类结果的堆叠条形图，调整横坐标标签的字体和角度 / Plot stacked bar chart of classification results
ggplot(combined_data, aes(x = Expression_Type, y = n, fill = Chosen_Expression)) +
  geom_bar(stat = "identity", position = "stack") +
  facet_wrap(~ Expressor) +
  labs(title = "Classification Results for Male29 and Male75",
       x = "Actual Expression",
       y = "Count of Predictions",
       fill = "Predicted Expression") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 10),
    axis.title.x = element_text(size = 12),
    axis.title.y = element_text(size = 12),
    plot.title = element_text(size = 14, face = "bold")
  )

# 安装并加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")
install.packages("ggplot2")
install.packages("stringr")  # 新增的包

# 加载所需的库
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)  # 新增的库

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型、expressor 和性别（Gender）
data <- data %>%
  mutate(Expression_Type = case_when(
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

# Step 4: 创建选择矩阵
uhr_results <- data %>%
  group_by(Expressor, Expression_Type, Material, Chosen_Expression, Gender) %>%
  summarise(n = n(), .groups = 'drop') %>%
  spread(Chosen_Expression, n, fill = 0) %>%
  ungroup()

# Step 5: 计算每个单元格的 UHR 值
uhr_results <- uhr_results %>%
  rowwise() %>%
  mutate(
    row_sum = sum(c_across(where(is.numeric)), na.rm = TRUE),
    across(
      where(is.numeric),
      ~ ifelse(row_sum == 0, 0, .x^2 / (row_sum * sum(.x))),
      .names = "UHR_{col}"
    )
  )

# Step 6: 使用子集选择所需的列
uhr_results_filtered <- uhr_results[, c("Expressor", "Gender", "Expression_Type", grep("UHR", names(uhr_results), value = TRUE))]

# Step 7: 计算每个 Expressor 在不同表情类型下的平均 UHR
uhr_by_expressor <- uhr_results_filtered %>%
  pivot_longer(cols = starts_with("UHR"), names_to = "Emotion", values_to = "UHR") %>%
  group_by(Expressor, Gender, Expression_Type) %>%
  summarise(Average_UHR = mean(UHR, na.rm = TRUE), .groups = 'drop')

# Step 8: 将表情类型缩写为前三个字母
uhr_by_expressor <- uhr_by_expressor %>%
  mutate(Expression_Type = substr(Expression_Type, 1, 3))

# Step 9: 保存每个 expressor 的 UHR 信息
write.csv(uhr_by_expressor, "UHR_by_Expressor_and_Emotion.csv", row.names = FALSE)

# Step 10: # 可视化 UHR 数据，并进一步调整横坐标标签和标题字体
ggplot(uhr_by_expressor, aes(x = Expression_Type, y = Average_UHR, fill = Gender)) +
  geom_bar(stat = "identity", position = "dodge") +
  facet_wrap(~ Expressor) +
  labs(title = "Average UHR by Expressor and Expression Type",
       x = "Expression Type",
       y = "Average UHR",
       fill = "Gender") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 60, hjust = 1, size = 5),  # 进一步缩小字体大小
    axis.title.x = element_text(size = 10),
    axis.title.y = element_text(size = 8),
    plot.title = element_text(size = 12, face = "bold"),  # 调整标题字体大小
    strip.text = element_text(size = 8)  # 缩小facet标签的字体大小
  )


# Step 11: 按性别分组并排序 expressors
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

# 检查最终排序结果
print("按平均 UHR 排序的 Female expressors：")
print(uhr_sorted_female)

print("按平均 UHR 排序的 Male expressors：")
print(uhr_sorted_male)

# Step 12: 将排序后的结果保存为 CSV 文件
write.csv(uhr_sorted_female, "Sorted_Female_Expressors_by_UHR.csv", row.names = FALSE)
write.csv(uhr_sorted_male, "Sorted_Male_Expressors_by_UHR.csv", row.names = FALSE)









# Step 4: 创建 Female88 和 Female26 的选择矩阵
female88_data <- data %>%
  filter(Expressor == "Fema88")

female26_data <- data %>%
  filter(Expressor == "Fema26")

female88_matrix <- female88_data %>%
  group_by(Expression_Type, Chosen_Expression) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  spread(Chosen_Expression, Count, fill = 0)

female26_matrix <- female26_data %>%
  group_by(Expression_Type, Chosen_Expression) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  spread(Chosen_Expression, Count, fill = 0)

# 将 Expression_Type 作为行名
rownames(female88_matrix) <- female88_matrix$Expression_Type
female88_matrix <- female88_matrix[ , -1]

rownames(female26_matrix) <- female26_matrix$Expression_Type
female26_matrix <- female26_matrix[ , -1]

# Step 6: 可视化
female88_long <- female88_data %>%
  group_by(Expression_Type, Chosen_Expression) %>%
  summarise(Count = n(), .groups = 'drop')

female26_long <- female26_data %>%
  group_by(Expression_Type, Chosen_Expression) %>%
  summarise(Count = n(), .groups = 'drop')

female88_long$Expressor <- "Female88"
female26_long$Expressor <- "Female26"

combined_data <- bind_rows(female88_long, female26_long)

# 绘制分类结果的堆叠条形图，调整横坐标标签的字体和角度
ggplot(combined_data, aes(x = Expression_Type, y = Count, fill = Chosen_Expression)) +
  geom_bar(stat = "identity", position = "stack") +
  facet_wrap(~ Expressor) +
  labs(title = "Classification Results for Female88 and Female26",
       x = "Actual Expression",
       y = "Count of Predictions",
       fill = "Predicted Expression") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 10), # 调整角度和大小
        axis.title.x = element_text(size = 12),
        axis.title.y = element_text(size = 12),
        plot.title = element_text(size = 14, face = "bold"))